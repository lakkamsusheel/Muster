'-------------------------------------------------------------------------------
' MUSTER.Info.FacilityInfo
'   P   ides the container to persist MUSTER Owner state
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        EN        12/04/04    Original class definition.
'  1.1        EN        12/29/04    Changed the datum,Method,LocationType Properties as integer from LookUpProperty Class.
'  1.2        AN        12/30/04    Added Try catch and Exception Handling/Logging
'  1.3        EN         1/04/05    taken out the default date values for datetransferred,datereceived,upcominginstalldate,datePowerOf variables
'  1.4        EN         1/10/05    Setting BolisDirty = false in reset, archive methods.
'  1.5        MNR       01/13/05    Added Events
'  1.6        AB        02/21/05    Added AgeThreshold and IsAgedData Attributes
'  1.7        MR        03/07/05    Changed Modified By and Modified On to read/write.
'  1.8        MR        03/14/05    Changed Created By and Created On to read/write.
'  1.9        MNR       03/15/05    Added Constructor New(ByVal drFacility As DataRow)
'  2.0        MNR       03/16/05    Removed strSrc from events
'  2.1        KKM       03/18/05    TankCollection and commentsCollection properties are added
'  2.2  Thomas Franey   06/16/09    Added Designated Operator for CAE Inspections Checklist
'
' Function          Description
'' New()             Instantiates an empty Facility object
''New(FACILITY_ID , FACILITY_AIID , ByVal NAME ,  OWNER_ID ,  ADDRESS_ID ,BILLING_ADDRESS_ID ,LATITUDE_DEGREE, _
'          LATITUDE_MINUTES, LATITUDE_SECONDS,  LONGITUDE_DEGREE , LONGITUDE_MINUTES , _
'          LONGITUDE_SECONDS , PHONE , FAX ,  FEES_PROFILE_ID ,  FACILITY_TYPE , _
'          FEES_STATUS ,  CURRENT_CIU_NUMBER , CAP_STATUS,  CAP_CANDIDATE ,  CITATION_PROFILE_ID, _
'          CURRENT_LUST_STATUS ,  FUEL_BRAND ,  FACILITY_DESCRIPTION ,SIGNATURE_NEEDED, DATE_RECD, _
'          DATE_TRANSFERRED ,  FACILITY_STATUS ,  DELETED , CREATED_BY ,DATE_CREATED ,LAST_EDITED_BY , _
'          DATE_LAST_EDITED ,  DATE_POWEROFF , UPCOMING_INSTALLATION , UPCOMING_INSTALLATION_DATE )
'                   Instantiates a populated Facility object
'Reset()           Sets the object state to the original state when loaded from or
'                   last saved to the repository
'Archive            Sets the object state to the old state when loaded from or
'                   last saved to the repository
'CheckDirty         'Check for dirty....
'Init               Intialise the object attributes...

' To Do Add attributes.... 
'Attribute          Description
' ID                The unique identifier associated with the Facility in the repository.
' Name              The name of the Facility.
' IsDirty           Indicates if the Facility state has been altered since it was
'                       last loaded from or saved to the repository.


'-------------------------------------------------------------------------------

Namespace MUSTER.Info
    <Serializable()> _
    Public Class FacilityInfo
#Region "Public Events"
        Public Event FacilityInfoChanged(ByVal bolValue As Boolean)
#End Region
#Region "Private member variables"
        Private nFacilityId As Integer
        Private strDesignatedOperator As String
        Private ostrDesignatedOperator As String
        Private strDesignatedManager As String
        Private ostrDesignatedManager As String
        Private nAIID As Integer
        Private strFacilityName As String
        Private strFacilityNameForEnsite As String
        Private nOwnerId As Integer
        Private nAddressId As Integer
        'Private strAddressLine1 As String
        'Private strAddressLine2 As String
        'Private strState As String
        'Private strCity As String
        'Private strZip As String
        'Private strFIPSCode As String
        Private nBillingAddId As Integer
        Private strLatitudeDegree As Single
        Private strLatitudeMin As Single
        Private strLatitudeSec As Double
        Private strLongitudeDegree As Single
        Private strLongitudeMin As Single
        Private strLongitudeSec As Double
        Private strPhone As String
        Private nDatum As Integer
        Private nMethod As Integer
        Private strFax As String
        Private nFEESPROFILEID As Integer
        Private nFacilityType As Integer
        Private nFeesStatus As Integer
        Private nCurrentCIUNumber As Integer
        Private nCapStatus As Integer
        Private bolCAPCandidate As Boolean
        Private nCitationProfileId As Integer
        Private nCurrentLUSTStatus As Integer
        Private strFuelBrand As String
        Private strFacilityDescription As String
        Private bolSignatureonNF As Boolean
        Private dtDateReceived As Date
        Private dtDateTransferred As Date
        Private nFacStatus As Integer
        Private bolDeleted As Boolean
        Private strCreatedBy As String
        Private dtCreatedOn As Date
        Private strModifiedBy As String
        Private dtModifiedOn As Date

        Private dtPowerOff As Date
        Private nLocationType As Integer
        Private bolUpcomingInstallation As Boolean
        Private dtUpcomingInstallationDate As Date
        Private bolIsDirty As Boolean
        Private nCurrentMGPTFStatus As String = String.Empty
        Private nLicenseeID As Integer
        Private nContractorID As Integer

        Private onFacilityId As Integer
        Private onAIID As Integer
        Private ostrFacilityName As String
        Private onCurrentMGPTFStatus As String = String.Empty
        Private ostrFacilityNameForEnsite As String
        Private onOwnerId As Integer
        Private onAddressId As Integer
        'Private ostrAddressLine1 As String
        'Private ostrAddressLine2 As String
        'Private ostrState As String
        'Private ostrCity As String
        'Private ostrZip As String
        'Private ostrFIPSCode As String
        Private onBillingAddId As Integer
        Private ostrLatitudeDegree As Single
        Private ostrLatitudeMin As Single
        Private ostrLatitudeSec As Double
        Private ostrLongitudeDegree As Single
        Private ostrLongitudeMin As Single
        Private ostrLongitudeSec As Double
        Private ostrPhone As String
        Private onDatum As Integer
        Private onMethod As Integer
        Private ostrFax As String
        Private onFEESPROFILEID As Integer
        Private onFacilityType As Integer
        Private onFeesStatus As Integer
        Private onCurrentCIUNumber As Integer
        Private onCapStatus As Integer
        Private obolCAPCandidate As Boolean
        Private onCitationProfileId As Integer
        Private onCurrentLUSTStatus As Integer
        Private ostrFuelBrand As String
        Private ostrFacilityDescription As String
        Private obolSignatureonNF As Boolean
        Private odtDateReceived As Date
        Private odtDateTransferred As Date
        Private onFacStatus As Integer
        Private obolDeleted As Boolean
        Private ostrCreatedBy As String
        Private odtCreatedOn As Date
        Private ostrModifiedBy As String
        Private odtModifiedOn As Date

        Private odtPowerOff As Date
        Private onLocationType As Integer
        Private obolUpcomingInstallation As Boolean
        Private odtUpcomingInstallationDate As Date
        Private obolIsDirty As Boolean
        Private onLicenseeID As Integer
        Private onContractorID As Integer



        Private dtDataAge As DateTime
        Private nAgeThreshold As Int16 = 5
        Private MusterException As New MUSTER.Exceptions.MusterExceptions

        Private colTank As MUSTER.Info.TankCollection
        'Private colComments As MUSTER.Info.CommentsCollection
        'Added by AB on 03/22/2005
        Private WithEvents colLustEvents As MUSTER.Info.LustEventCollection
        Private colClosureEvent As MUSTER.Info.ClosureEventCollection
        'End changes

#End Region
#Region "Constructors"
        Sub New()
            MyBase.New()
            dtDataAge = Now()
            Me.Init()
            Me.InitCollections()
        End Sub
        ' The prototype New method
        Public Sub New(ByVal FACILITY_ID As Integer, ByVal FACILITY_AIID As Integer, ByVal NAME As String, _
            ByVal OWNER_ID As Integer, _
            ByVal ADDRESS_ID As Integer, _
            ByVal BILLING_ADDRESS_ID As Integer, ByVal LATITUDE_DEGREE As Single, _
            ByVal LATITUDE_MINUTES As Single, ByVal LATITUDE_SECONDS As Double, ByVal LONGITUDE_DEGREE As Single, ByVal LONGITUDE_MINUTES As Single, _
            ByVal LONGITUDE_SECONDS As Double, ByVal PHONE As String, ByVal DATUM As Integer, ByVal METHOD As Integer, ByVal FAX As String, ByVal FEES_PROFILE_ID As Integer, ByVal FACILITY_TYPE As Single, _
            ByVal FEES_STATUS As Integer, ByVal CURRENT_CIU_NUMBER As Integer, ByVal CAP_STATUS As Integer, ByVal CAP_CANDIDATE As Boolean, ByVal CITATION_PROFILE_ID As Integer, _
            ByVal CURRENT_LUST_STATUS As Integer, ByVal FUEL_BRAND As String, ByVal FACILITY_DESCRIPTION As String, ByVal SIGNATURE_NEEDED As Boolean, ByVal DATE_RECD As Date, _
            ByVal DATE_TRANSFERRED As Date, ByVal FACILITY_STATUS As Integer, ByVal DELETED As Boolean, ByVal CREATED_BY As String, ByVal DATE_CREATED As Date, ByVal LAST_EDITED_BY As String, _
            ByVal DATE_LAST_EDITED As Date, ByVal DATE_POWEROFF As Date, ByVal LocationType As Integer, ByVal UPCOMING_INSTALLATION As Boolean, ByVal UPCOMING_INSTALLATION_DATE As Date, _
            ByVal LICENSEE_ID As Integer, ByVal CONTRACTOR_ID As Integer, ByVal NameForEnsite As String, Optional ByVal MGPTF As String = "", Optional ByVal ModifiedOnModule As String = "", Optional ByVal ModifiedByModule As String = "", _
            Optional ByVal DESIGNATED_OPERATOR As String = "", _
            Optional ByVal DESIGNATED_MANAGER As String = "")

            'ByVal AddressLineOne As String, _
            'ByVal AddressTwo As String, _
            'ByVal City As String, _
            'ByVal State As String, _
            'ByVal Zip As String, _
            'ByVal FIPSCode As String, _

            onFacilityId = FACILITY_ID
            onAIID = FACILITY_AIID
            ostrFacilityName = NAME
            ostrFacilityNameForEnsite = NameForEnsite
            onOwnerId = OWNER_ID
            onAddressId = ADDRESS_ID
            'ostrAddressLine1 = AddressLineOne
            'ostrAddressLine2 = AddressTwo
            'ostrCity = City
            'ostrState = State
            'ostrZip = Zip
            'ostrFIPSCode = FIPSCode
            onBillingAddId = BILLING_ADDRESS_ID
            ostrLatitudeDegree = LATITUDE_DEGREE
            ostrLatitudeMin = LATITUDE_MINUTES
            ostrLatitudeSec = LATITUDE_SECONDS
            ostrLongitudeDegree = LONGITUDE_DEGREE
            ostrLongitudeMin = LONGITUDE_MINUTES
            ostrLongitudeSec = LONGITUDE_SECONDS
            ostrDesignatedOperator = DESIGNATED_OPERATOR
            ostrDesignatedManager = DESIGNATED_MANAGER
            ostrPhone = PHONE
            onDatum = DATUM
            onMethod = METHOD
            ostrFax = FAX
            onFEESPROFILEID = FEES_PROFILE_ID
            onFacilityType = FACILITY_TYPE
            onFeesStatus = FEES_STATUS
            onCurrentCIUNumber = CURRENT_CIU_NUMBER
            onCapStatus = CAP_STATUS
            obolCAPCandidate = CAP_CANDIDATE
            onCitationProfileId = CITATION_PROFILE_ID
            onCurrentLUSTStatus = CURRENT_LUST_STATUS
            ostrFuelBrand = FUEL_BRAND
            ostrFacilityDescription = FACILITY_DESCRIPTION
            obolSignatureonNF = SIGNATURE_NEEDED
            odtDateReceived = DATE_RECD.Date
            odtDateTransferred = DATE_TRANSFERRED.Date
            onFacStatus = FACILITY_STATUS
            obolDeleted = DELETED
            ostrCreatedBy = CREATED_BY
            odtCreatedOn = DATE_CREATED
            ostrModifiedBy = LAST_EDITED_BY
            odtModifiedOn = DATE_LAST_EDITED
            odtPowerOff = DATE_POWEROFF.Date
            onLocationType = LocationType
            obolUpcomingInstallation = UPCOMING_INSTALLATION
            odtUpcomingInstallationDate = UPCOMING_INSTALLATION_DATE.Date
            onLicenseeID = LICENSEE_ID
            onContractorID = CONTRACTOR_ID
            onCurrentMGPTFStatus = MGPTF
            dtDataAge = Now()


            Me.InitCollections()
            Me.Reset()
        End Sub
        Sub New(ByVal drFacility As DataRow)
            Try
                onFacilityId = drFacility.Item("FACILITY_ID")
                onAIID = IIf(drFacility.Item("FACILITY_AIID") Is System.DBNull.Value, 0, drFacility.Item("FACILITY_AIID"))
                ostrFacilityName = IIf(drFacility.Item("NAME") Is System.DBNull.Value, String.Empty, drFacility.Item("NAME"))
                ostrFacilityNameForEnsite = IIf(drFacility.Item("NAME_FOR_ENSITE") Is System.DBNull.Value, String.Empty, drFacility.Item("NAME_FOR_ENSITE"))
                onOwnerId = IIf(drFacility.Item("OWNER_ID") Is System.DBNull.Value, 0, drFacility.Item("OWNER_ID"))
                onAddressId = IIf(drFacility.Item("ADDRESS_ID") Is System.DBNull.Value, 0, drFacility.Item("ADDRESS_ID"))
                'ostrAddressLine1 = drFacility.Item("ADDRESS_LINE_ONE")
                'ostrAddressLine2 = IIf(drFacility.Item("ADDRESS_TWO") Is System.DBNull.Value, String.Empty, drFacility.Item("ADDRESS_TWO"))
                'ostrCity = IIf(drFacility.Item("CITY") Is System.DBNull.Value, String.Empty, drFacility.Item("CITY"))
                'ostrState = IIf(drFacility.Item("STATE") Is System.DBNull.Value, String.Empty, drFacility.Item("STATE"))
                'ostrZip = IIf(drFacility.Item("ZIP") Is System.DBNull.Value, String.Empty, drFacility.Item("ZIP"))
                'ostrFIPSCode = IIf(drFacility.Item("FIPS_CODE") Is System.DBNull.Value, String.Empty, drFacility.Item("FIPS_CODE"))
                onBillingAddId = IIf(drFacility.Item("BILLING_ADDRESS_ID") Is System.DBNull.Value, 0, drFacility.Item("BILLING_ADDRESS_ID"))
                ostrLatitudeDegree = IIf(drFacility.Item("LATITUDE_DEGREE") Is System.DBNull.Value, -1, drFacility.Item("LATITUDE_DEGREE"))
                ostrLatitudeMin = IIf(drFacility.Item("LATITUDE_MINUTES") Is System.DBNull.Value, -1, drFacility.Item("LATITUDE_MINUTES"))
                ostrLatitudeSec = IIf(drFacility.Item("LATITUDE_SECONDS") Is System.DBNull.Value, -1, drFacility.Item("LATITUDE_SECONDS"))
                ostrLongitudeDegree = IIf(drFacility.Item("LONGITUDE_DEGREE") Is System.DBNull.Value, -1, drFacility.Item("LONGITUDE_DEGREE"))
                ostrLongitudeMin = IIf(drFacility.Item("LONGITUDE_MINUTES") Is System.DBNull.Value, -1, drFacility.Item("LONGITUDE_MINUTES"))
                ostrLongitudeSec = IIf(drFacility.Item("LONGITUDE_SECONDS") Is System.DBNull.Value, -1, drFacility.Item("LONGITUDE_SECONDS"))
                ostrPhone = IIf(drFacility.Item("PHONE") Is System.DBNull.Value, String.Empty, drFacility.Item("PHONE"))
                onDatum = IIf(drFacility.Item("DATUM") Is System.DBNull.Value, 0, drFacility.Item("DATUM"))
                onMethod = IIf(drFacility.Item("METHOD") Is System.DBNull.Value, 0, drFacility.Item("METHOD"))
                ostrFax = IIf(drFacility.Item("FAX") Is System.DBNull.Value, String.Empty, drFacility.Item("FAX"))
                onFEESPROFILEID = IIf(drFacility.Item("FEES_PROFILE_ID") Is System.DBNull.Value, 0, drFacility.Item("FEES_PROFILE_ID"))
                onFacilityType = IIf(drFacility.Item("FACILITY_TYPE") Is System.DBNull.Value, 0, drFacility.Item("FACILITY_TYPE"))
                onFeesStatus = IIf(drFacility.Item("FEES_STATUS") Is System.DBNull.Value, 0, drFacility.Item("FEES_STATUS"))
                onCurrentCIUNumber = IIf(drFacility.Item("CURRENT_CIU_NUMBER") Is System.DBNull.Value, 0, drFacility.Item("CURRENT_CIU_NUMBER"))
                onCapStatus = IIf(drFacility.Item("CAP_STATUS") Is System.DBNull.Value, 0, drFacility.Item("CAP_STATUS"))
                obolCAPCandidate = IIf(drFacility.Item("CAP_CANDIDATE") Is System.DBNull.Value, False, drFacility.Item("CAP_CANDIDATE"))
                onCitationProfileId = IIf(drFacility.Item("CITATION_PROFILE_ID") Is System.DBNull.Value, 0, drFacility.Item("CITATION_PROFILE_ID"))
                onCurrentLUSTStatus = IIf(drFacility.Item("CURRENT_LUST_STATUS") Is System.DBNull.Value, 0, drFacility.Item("CURRENT_LUST_STATUS"))
                ostrFuelBrand = IIf(drFacility.Item("FUEL_BRAND") Is System.DBNull.Value, String.Empty, drFacility.Item("FUEL_BRAND"))
                ostrFacilityDescription = IIf(drFacility.Item("FACILITY_DESCRIPTION") Is System.DBNull.Value, String.Empty, drFacility.Item("FACILITY_DESCRIPTION"))
                obolSignatureonNF = IIf(drFacility.Item("SIGNATURE_NEEDED") Is System.DBNull.Value, False, drFacility.Item("SIGNATURE_NEEDED"))
                odtDateReceived = IIf(drFacility.Item("DATE_RECD") Is System.DBNull.Value, CDate("01/01/0001"), drFacility.Item("DATE_RECD"))
                odtDateReceived = odtDateReceived.Date
                odtDateTransferred = IIf(drFacility.Item("DATE_TRANSFERRED") Is System.DBNull.Value, CDate("01/01/0001"), drFacility.Item("DATE_TRANSFERRED"))
                odtDateTransferred = odtDateTransferred.Date
                onFacStatus = IIf(drFacility.Item("FACILITY_STATUS") Is System.DBNull.Value, 515, drFacility.Item("FACILITY_STATUS"))
                obolDeleted = IIf(drFacility.Item("DELETED") Is System.DBNull.Value, False, drFacility.Item("DELETED"))
                ostrCreatedBy = IIf(drFacility.Item("CREATED_BY") Is System.DBNull.Value, String.Empty, drFacility.Item("CREATED_BY"))
                odtCreatedOn = IIf(drFacility.Item("DATE_CREATED") Is System.DBNull.Value, CDate("01/01/0001"), drFacility.Item("DATE_CREATED"))
                ostrModifiedBy = IIf(drFacility.Item("LAST_EDITED_BY") Is System.DBNull.Value, String.Empty, drFacility.Item("LAST_EDITED_BY"))
                odtModifiedOn = IIf(drFacility.Item("DATE_LAST_EDITED") Is System.DBNull.Value, CDate("01/01/0001"), drFacility.Item("DATE_LAST_EDITED"))
                odtPowerOff = IIf(drFacility.Item("DATE_POWEROFF") Is System.DBNull.Value, CDate("01/01/0001"), drFacility.Item("DATE_POWEROFF"))
                odtPowerOff = odtPowerOff.Date
                onLocationType = IIf(drFacility.Item("LOCATION_TYPE") Is System.DBNull.Value, 0, drFacility.Item("LOCATION_TYPE"))
                obolUpcomingInstallation = IIf(drFacility.Item("UPCOMING_INSTALLATION") Is System.DBNull.Value, False, drFacility.Item("UPCOMING_INSTALLATION"))
                odtUpcomingInstallationDate = IIf(drFacility.Item("UPCOMING_INSTALLATION_DATE") Is System.DBNull.Value, CDate("01/01/0001"), drFacility.Item("UPCOMING_INSTALLATION_DATE"))
                odtUpcomingInstallationDate = odtUpcomingInstallationDate.Date
                onLicenseeID = IIf(drFacility.Item("LICENSEEID") Is System.DBNull.Value, 0, drFacility.Item("LICENSEEID"))
                onContractorID = IIf(drFacility.Item("CONTRACTORID") Is System.DBNull.Value, 0, drFacility.Item("CONTRACTORID"))
                onCurrentMGPTFStatus = IIf(drFacility.Item("MGPTFStatus") Is DBNull.Value, String.Empty, drFacility.Item("MGPTFStatus"))
                ostrDesignatedOperator = IIf(drFacility("DesignatedOperator") Is DBNull.Value, String.Empty, drFacility.Item("DesignatedOperator"))
                ostrDesignatedManager = IIf(drFacility("DesignatedManager") Is DBNull.Value, String.Empty, drFacility.Item("DesignatedManager"))



                Me.InitCollections()
                dtDataAge = Now()
                Me.Reset()
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
#End Region
#Region "Exposed Operations"
        Public Sub Reset()
            If nFacilityId >= 0 Then
                nFacilityId = onFacilityId
            End If
            nAIID = onAIID
            strFacilityName = ostrFacilityName
            strFacilityNameForEnsite = ostrFacilityNameForEnsite
            nOwnerId = onOwnerId
            nAddressId = onAddressId
            'strAddressLine1 = ostrAddressLine1
            'strAddressLine2 = ostrAddressLine2
            'strState = ostrState
            'strCity = ostrCity
            'strZip = ostrZip
            'strFIPSCode = ostrFIPSCode
            nBillingAddId = onBillingAddId
            strLatitudeDegree = ostrLatitudeDegree
            strLatitudeMin = ostrLatitudeMin
            strLatitudeSec = ostrLatitudeSec
            strLongitudeDegree = ostrLongitudeDegree
            strLongitudeMin = ostrLongitudeMin
            strLongitudeSec = ostrLongitudeSec
            strPhone = ostrPhone
            strFax = ostrFax
            nFEESPROFILEID = onFEESPROFILEID
            nFacilityType = onFacilityType
            strDesignatedOperator = ostrDesignatedOperator
            strDesignatedManager = ostrDesignatedManager
            nFeesStatus = onFeesStatus
            nCurrentCIUNumber = onCurrentCIUNumber
            nCapStatus = onCapStatus
            bolCAPCandidate = obolCAPCandidate
            nCitationProfileId = onCitationProfileId
            nCurrentLUSTStatus = onCurrentLUSTStatus
            strFuelBrand = ostrFuelBrand
            strFacilityDescription = ostrFacilityDescription
            bolSignatureonNF = obolSignatureonNF
            bolDeleted = obolDeleted
            strCreatedBy = ostrCreatedBy
            dtCreatedOn = odtCreatedOn
            strModifiedBy = ostrModifiedBy
            dtModifiedOn = odtModifiedOn

            dtDateReceived = odtDateReceived
            dtDateTransferred = odtDateTransferred
            nFacStatus = onFacStatus
            nDatum = onDatum
            nMethod = onMethod
            nLocationType = onLocationType
            dtPowerOff = odtPowerOff
            bolUpcomingInstallation = obolUpcomingInstallation
            dtUpcomingInstallationDate = odtUpcomingInstallationDate
            nLicenseeID = onLicenseeID
            nContractorID = onContractorID
            nCurrentMGPTFStatus = onCurrentMGPTFStatus
            bolIsDirty = False
            RaiseEvent FacilityInfoChanged(bolIsDirty)
        End Sub
        Public Sub Archive()
            onFacilityId = nFacilityId
            onAIID = nAIID
            ostrFacilityName = strFacilityName
            ostrFacilityNameForEnsite = strFacilityNameForEnsite
            onOwnerId = nOwnerId
            onAddressId = nAddressId
            'ostrAddressLine1 = strAddressLine1
            'ostrAddressLine2 = strAddressLine2
            'ostrState = strState
            'ostrCity = strCity
            'ostrZip = strZip
            'ostrFIPSCode = strFIPSCode
            onBillingAddId = nBillingAddId
            ostrLatitudeDegree = strLatitudeDegree
            ostrLatitudeMin = strLatitudeMin
            ostrLatitudeSec = strLatitudeSec
            ostrLongitudeDegree = strLongitudeDegree
            ostrLongitudeMin = strLongitudeMin
            ostrLongitudeSec = strLongitudeSec
            ostrPhone = strPhone
            ostrFax = strFax
            onFEESPROFILEID = nFEESPROFILEID
            onFacilityType = nFacilityType
            onFeesStatus = nFeesStatus
            onCurrentCIUNumber = nCurrentCIUNumber
            onCapStatus = nCapStatus
            obolCAPCandidate = bolCAPCandidate
            onCitationProfileId = nCitationProfileId
            onCurrentLUSTStatus = nCurrentLUSTStatus
            ostrFuelBrand = strFuelBrand
            ostrDesignatedOperator = strDesignatedOperator
            ostrDesignatedManager = strDesignatedManager
            onCurrentMGPTFStatus = nCurrentMGPTFStatus
            ostrFacilityDescription = strFacilityDescription
            obolSignatureonNF = bolSignatureonNF
            obolDeleted = bolDeleted
            ostrCreatedBy = strCreatedBy
            odtCreatedOn = dtCreatedOn
            ostrModifiedBy = strModifiedBy
            odtModifiedOn = dtModifiedOn
            odtDateReceived = dtDateReceived
            odtDateTransferred = dtDateTransferred
            onFacStatus = nFacStatus
            onDatum = nDatum
            onMethod = nMethod
            onLocationType = nLocationType
            odtPowerOff = dtPowerOff
            obolUpcomingInstallation = bolUpcomingInstallation
            odtUpcomingInstallationDate = dtUpcomingInstallationDate
            onLicenseeID = nLicenseeID
            onContractorID = nContractorID
            bolIsDirty = False


        End Sub
#End Region
#Region "Private Operations"
        Private Sub CheckDirty()
            Dim obolIsDirty As Boolean = bolIsDirty
            'strAddressLine1 <> ostrAddressLine1 Or _
            'strAddressLine2 <> ostrAddressLine2 Or _
            'strState <> ostrState Or _
            'strCity <> ostrCity Or _
            'strZip <> ostrZip Or _
            'strFIPSCode <> ostrFIPSCode Or _
            '(nCapStatus <> onCapStatus) Or _
            bolIsDirty = (nAIID <> onAIID) Or _
                (strFacilityName <> ostrFacilityName) Or _
                (strFacilityNameForEnsite <> ostrFacilityNameForEnsite) Or _
                (nOwnerId <> onOwnerId) Or _
                (nAddressId <> onAddressId) Or _
                (nBillingAddId <> onBillingAddId) Or _
                (strLatitudeDegree <> ostrLatitudeDegree) Or _
                (strLatitudeMin <> ostrLatitudeMin) Or _
                (strLatitudeSec <> ostrLatitudeSec) Or _
                (strLongitudeDegree <> ostrLongitudeDegree) Or _
                (strLongitudeMin <> ostrLongitudeMin) Or _
                (strLongitudeSec <> ostrLongitudeSec) Or _
                (strPhone <> ostrPhone) Or _
                (strFax <> ostrFax) Or _
                (nFEESPROFILEID <> onFEESPROFILEID) Or _
                (nFacilityType <> onFacilityType) Or _
                (nFeesStatus <> onFeesStatus) Or _
                (nCurrentCIUNumber <> onCurrentCIUNumber) Or _
                (bolCAPCandidate <> obolCAPCandidate) Or _
                (nCitationProfileId <> onCitationProfileId) Or _
                (strFuelBrand <> ostrFuelBrand) Or _
                (strFacilityDescription <> ostrFacilityDescription) Or _
                (bolSignatureonNF <> obolSignatureonNF) Or _
                (bolDeleted <> obolDeleted) Or _
                (dtDateReceived <> odtDateReceived) Or _
                (dtDateTransferred <> odtDateTransferred) Or _
                (nFacStatus <> onFacStatus) Or _
                (dtPowerOff <> odtPowerOff) Or _
                (bolUpcomingInstallation <> obolUpcomingInstallation) Or _
                (dtUpcomingInstallationDate <> odtUpcomingInstallationDate) Or _
                (nDatum <> onDatum) Or _
                (nMethod <> onMethod) Or _
                (nLocationType <> onLocationType) Or _
                (nLicenseeID <> onLicenseeID) Or _
                (nContractorID <> onContractorID) Or _
                (strDesignatedOperator <> ostrDesignatedOperator) Or _
                (strDesignatedManager <> ostrDesignatedManager)



            If obolIsDirty <> bolIsDirty Then
                'MsgBox("Info F:" + ID.ToString)
                RaiseEvent FacilityInfoChanged(bolIsDirty)
            End If
        End Sub
        Private Sub Init()
            'obolIsNewItem = False
            onFacilityId = 0
            onAIID = 0
            onOwnerId = 0
            onAddressId = 0
            'ostrAddressLine1 = String.Empty
            'ostrAddressLine2 = String.Empty
            'ostrCity = String.Empty
            'ostrState = String.Empty
            'ostrZip = String.Empty
            'ostrFIPSCode = String.Empty
            onBillingAddId = 0
            ostrFacilityName = String.Empty
            ostrFacilityNameForEnsite = String.Empty
            ostrLatitudeDegree = -1.0
            ostrLatitudeMin = -1.0
            ostrLatitudeSec = -1.0
            ostrLongitudeDegree = -1.0
            ostrLongitudeMin = -1.0
            ostrDesignatedOperator = String.Empty
            ostrDesignatedManager = String.Empty
            ostrLongitudeSec = -1.0
            ostrPhone = String.Empty
            ostrFax = String.Empty
            onFEESPROFILEID = 0
            onFacilityType = 0
            onFeesStatus = 0
            onCurrentCIUNumber = 0
            onCapStatus = 0
            obolCAPCandidate = False
            onCitationProfileId = 0
            onCurrentLUSTStatus = 0
            ostrFuelBrand = String.Empty
            ostrFacilityDescription = String.Empty
            obolSignatureonNF = False
            obolDeleted = False
            ostrCreatedBy = String.Empty
            odtDateReceived = Nothing
            odtUpcomingInstallationDate = Nothing
            odtCreatedOn = CDate("01/01/0001")
            ostrModifiedBy = String.Empty
            odtModifiedOn = CDate("01/01/0001")
            odtDateTransferred = Nothing
            onFacStatus = 515 'Closed
            odtPowerOff = Nothing
            obolUpcomingInstallation = False
            onDatum = 0
            onMethod = 0
            onLocationType = 0
            onLicenseeID = 0
            onContractorID = 0
            onCurrentMGPTFStatus = String.Empty
            obolIsDirty = False


            Me.Reset()
        End Sub
        Private Sub InitCollections()
            colTank = New MUSTER.Info.TankCollection
            'colComments = New MUSTER.Info.CommentsCollection
            colLustEvents = New MUSTER.Info.LustEventCollection
            colClosureEvent = New MUSTER.Info.ClosureEventCollection
        End Sub
#End Region
#Region "Exposed Attributes"

       

        Public ReadOnly Property CurrentMGPTFStatus() As String
            Get
                Return Me.nCurrentMGPTFStatus
            End Get
        End Property

        Public Property LustEventCollection() As MUSTER.Info.LustEventCollection
            Get
                Return colLustEvents
            End Get
            Set(ByVal Value As MUSTER.Info.LustEventCollection)
                colLustEvents = Value
            End Set
        End Property

        Public Property ClosureEventCollection() As MUSTER.Info.ClosureEventCollection
            Get
                Return colClosureEvent
            End Get
            Set(ByVal Value As MUSTER.Info.ClosureEventCollection)
                colClosureEvent = Value
            End Set
        End Property
        Public Property TankCollection() As MUSTER.Info.TankCollection
            Get
                Return colTank
            End Get
            Set(ByVal Value As MUSTER.Info.TankCollection)
                colTank = Value
            End Set
        End Property
        'Public Property CommentsCollection() As MUSTER.Info.CommentsCollection
        '    Get
        '        Return colComments
        '    End Get
        '    Set(ByVal Value As MUSTER.Info.CommentsCollection)
        '        colComments = Value
        '    End Set
        'End Property
        Public Property ID() As Integer

            Get
                Return Me.nFacilityId
            End Get
            Set(ByVal value As Integer)
                Me.nFacilityId = value
                Me.CheckDirty()
            End Set
        End Property
        Public Property AIID() As Integer

            Get
                Return Me.nAIID
            End Get
            Set(ByVal value As Integer)
                Me.nAIID = value
                Me.CheckDirty()
            End Set
        End Property
        Public Property Name() As String
            Get
                Return Me.strFacilityName
            End Get
            Set(ByVal value As String)
                Me.strFacilityName = value
                Me.CheckDirty()
            End Set
        End Property
        Public Property NameForEnsite() As String
            Get
                Return strFacilityNameForEnsite
            End Get
            Set(ByVal Value As String)
                strFacilityNameForEnsite = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property OwnerID() As Integer

            Get
                Return Me.nOwnerId
            End Get
            Set(ByVal value As Integer)
                Me.nOwnerId = value
                Me.CheckDirty()
            End Set
        End Property
        Public Property AddressID() As Integer
            Get
                Return Me.nAddressId
            End Get
            Set(ByVal value As Integer)
                Me.nAddressId = value
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
        Public Property BillingAddressID() As Integer
            Get
                Return Me.nBillingAddId
            End Get
            Set(ByVal value As Integer)
                Me.nBillingAddId = value
                Me.CheckDirty()
            End Set
        End Property
        Public Property LatitudeDegree() As Single

            Get
                Return Me.strLatitudeDegree
            End Get
            Set(ByVal value As Single)
                Me.strLatitudeDegree = value
                Me.CheckDirty()
            End Set
        End Property
        Public Property LatitudeMinutes() As Single
            Get
                Return Me.strLatitudeMin
            End Get
            Set(ByVal value As Single)
                Me.strLatitudeMin = value
                Me.CheckDirty()
            End Set
        End Property
        Public Property LatitudeSeconds() As Double
            Get
                Return Me.strLatitudeSec
            End Get
            Set(ByVal value As Double)
                Me.strLatitudeSec = value
                Me.CheckDirty()
            End Set
        End Property
        Public Property LongitudeDegree() As Single
            Get
                Return Me.strLongitudeDegree
            End Get
            Set(ByVal value As Single)
                Me.strLongitudeDegree = value
                Me.CheckDirty()
            End Set
        End Property
        Public Property LongitudeMinutes() As Single
            Get
                Return Me.strLongitudeMin
            End Get
            Set(ByVal value As Single)
                Me.strLongitudeMin = value
                Me.CheckDirty()
            End Set
        End Property
        Public Property LongitudeSeconds() As Double

            Get
                Return Me.strLongitudeSec
            End Get
            Set(ByVal value As Double)
                Me.strLongitudeSec = value
                Me.CheckDirty()
            End Set
        End Property
        Public Property Phone() As String

            Get
                Return Me.strPhone
            End Get
            Set(ByVal value As String)
                Me.strPhone = value
                Me.CheckDirty()
            End Set
        End Property
        Public Property Datum() As Integer
            Get
                Return nDatum
            End Get
            Set(ByVal Value As Integer)
                nDatum = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property Method() As Integer
            Get
                Return nMethod
            End Get
            Set(ByVal Value As Integer)
                nMethod = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property Fax() As String

            Get
                Return Me.strFax
            End Get
            Set(ByVal value As String)
                Me.strFax = value
                Me.CheckDirty()
            End Set
        End Property
        Public Property FeesProfileId() As Integer

            Get
                Return Me.nFEESPROFILEID
            End Get
            Set(ByVal value As Integer)
                Me.nFEESPROFILEID = value
                Me.CheckDirty()
            End Set
        End Property
        Public Property FacilityType() As Integer
            Get
                Return Me.nFacilityType
            End Get
            Set(ByVal value As Integer)
                Me.nFacilityType = value
                Me.CheckDirty()
            End Set
        End Property
        Public Property FeesStatus() As Integer

            Get
                Return Me.nFeesStatus
            End Get
            Set(ByVal value As Integer)
                Me.nFeesStatus = value
                Me.CheckDirty()
            End Set
        End Property
        Public Property CurrentCIUNumber() As Integer

            Get
                Return Me.nCurrentCIUNumber
            End Get
            Set(ByVal value As Integer)
                Me.nCurrentCIUNumber = value
                Me.CheckDirty()
            End Set
        End Property
        Public Property CapStatus() As Integer
            Get
                Return nCapStatus
            End Get
            Set(ByVal value As Integer)
                Me.nCapStatus = value
                Me.CheckDirty()
            End Set
        End Property
        Public Property CapStatusOriginal() As Integer
            Get
                Return onCapStatus
            End Get
            Set(ByVal value As Integer)
                onCapStatus = value
                Me.CheckDirty()
            End Set
        End Property
        Public Property CAPCandidate() As Boolean
            Get
                Return bolCAPCandidate
            End Get
            Set(ByVal Value As Boolean)
                bolCAPCandidate = Value
                Me.CheckDirty()
            End Set
        End Property
        Public ReadOnly Property CAPCandidateOriginal() As Boolean
            Get
                Return obolCAPCandidate
            End Get
        End Property
        Public Property CitationProfileID() As Integer

            Get
                Return Me.nCitationProfileId
            End Get
            Set(ByVal value As Integer)
                Me.nCitationProfileId = value
                Me.CheckDirty()
            End Set
        End Property
        Public Property CurrentLUSTStatus() As Integer

            Get
                Return Me.nCurrentLUSTStatus
            End Get
            Set(ByVal value As Integer)
                Me.nCurrentLUSTStatus = value
                'Me.CheckDirty()
            End Set
        End Property
        Public Property FuelBrand() As String
            Get
                Return strFuelBrand
            End Get
            Set(ByVal Value As String)
                strFuelBrand = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property FacilityDescription() As String

            Get
                Return Me.strFacilityDescription
            End Get
            Set(ByVal value As String)
                Me.strFacilityDescription = value
                Me.CheckDirty()
            End Set
        End Property
        Public Property SignatureOnNF() As Boolean
            Get
                Return bolSignatureonNF
            End Get
            Set(ByVal Value As Boolean)
                bolSignatureonNF = Value
                Me.CheckDirty()
            End Set
        End Property
        Public ReadOnly Property SignatureOnNFOriginal() As Boolean
            Get
                Return obolSignatureonNF
            End Get
        End Property
        Public Property DateReceived() As Date
            Get
                Return dtDateReceived.Date
            End Get
            Set(ByVal Value As Date)
                dtDateReceived = Value.Date
                Me.CheckDirty()
            End Set
        End Property
        Public Property DateTransferred() As Date
            Get
                Return dtDateTransferred.Date
            End Get
            Set(ByVal Value As Date)
                dtDateTransferred = Value.Date
                Me.CheckDirty()
            End Set
        End Property
        Public Property FacilityStatus() As Integer
            Get
                Return nFacStatus
            End Get
            Set(ByVal Value As Integer)
                nFacStatus = Value
                Me.CheckDirty()
            End Set
        End Property
        Public WriteOnly Property FacilityStatusOriginal() As Integer
            Set(ByVal Value As Integer)
                onFacStatus = Value
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
        Public Property DatePowerOff() As Date
            Get
                Return dtPowerOff.Date
            End Get
            Set(ByVal Value As Date)
                dtPowerOff = Value.Date
                Me.CheckDirty()
            End Set
        End Property
        Public Property LocationType() As Integer
            Get
                Return nLocationType
            End Get
            Set(ByVal Value As Integer)
                nLocationType = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property UpcomingInstallation() As Boolean
            Get
                Return bolUpcomingInstallation
            End Get
            Set(ByVal Value As Boolean)
                bolUpcomingInstallation = Value
                Me.CheckDirty()
            End Set
        End Property
        Public ReadOnly Property UpcomingInstallationOriginal() As Boolean
            Get
                Return obolUpcomingInstallation
            End Get
        End Property
        Public Property UpcomingInstallationDate() As Date
            Get
                Return dtUpcomingInstallationDate.Date
            End Get
            Set(ByVal Value As Date)
                dtUpcomingInstallationDate = Value.Date
                Me.CheckDirty()
            End Set
        End Property
        Public ReadOnly Property UpcomingInstallationDateOriginal() As Date
            Get
                Return odtUpcomingInstallationDate.Date
            End Get
        End Property
        Public Property LicenseeID() As Integer
            Get
                Return nLicenseeID
            End Get
            Set(ByVal Value As Integer)
                nLicenseeID = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property ContractorID() As Integer
            Get
                Return nContractorID
            End Get
            Set(ByVal Value As Integer)
                nContractorID = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property IsDirty() As Boolean
            Get
                If bolIsDirty Then Return bolIsDirty
                For Each tank As MUSTER.Info.TankInfo In TankCollection.Values
                    If tank.IsDirty Then
                        Return True
                    End If
                Next
                Return False
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
        Public ReadOnly Property ChildrenDirty() As Boolean
            Get
                For Each tnk As MUSTER.Info.TankInfo In TankCollection.Values
                    If tnk.IsDirty Or tnk.ChildrenDirty Then Return True
                Next
                Return False
            End Get
        End Property
#End Region
#Region "iAccessors"

        Public Property CreatedBy() As String
            Get
                Return strCreatedBy
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
                Return strModifiedBy
            End Get
            Set(ByVal Value As String)
                strModifiedBy = Value
            End Set
        End Property

        Public Property DesignatedOperator() As String
            Get
                Return strDesignatedOperator
            End Get
            Set(ByVal Value As String)
                strDesignatedOperator = Value
            End Set
        End Property
        Public Property DesignatedManager() As String
            Get
                Return strDesignatedManager
            End Get
            Set(ByVal Value As String)
                strDesignatedManager = Value
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

