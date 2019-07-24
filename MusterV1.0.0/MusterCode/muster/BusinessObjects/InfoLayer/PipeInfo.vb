'-------------------------------------------------------------------------------
' MUSTER.Info.PipeInfo
'   Provides the container to persist MUSTER Pipe state
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        MNR       12/15/04    Original class definition
'  1.1        MNR       12/22/04    Added function Validate(..) in save function
'  1.2        AN        12/30/04    Added Try catch and Exception Handling/Logging
'  1.3        MNR        1/04/05    Deleted function Validate(..) - transferred to Pipe.vb
'  1.4        EN         1/06/05    Added Events and Raised the Events..
'  1.5        EN         1/19/05    Added Source Column in the Event.
'  1.6        MR        03/07/05    Change Modified By and Modified On to read/write.
'  1.7        MNR       03/10/05    Added FacilityPowerOff Property
'  1.8        MR        03/14/05    Change Created By and Created On to read/write.
'  1.9        MNR       03/15/05    Updated Constructor New(ByVal drPipe As DataRow) to check for System.DBNull.Value
'  2.0        AB        03/16/05    Added AgeThreshold and IsAgedData Attributes
'  2.1        MNR       03/16/05    Removed strSrc from events
'  2.2        KKM       03/18/05    CommentsCollection property is added
'  2.3   Thomas Franey  02/23/09    Added Parent pipe ID field to all areas needed

'
' Function          Description
' New()             Instantiates an empty PipeInfo object
' New(...)          Instantiates a populated PipeInfo object
' Reset()           Sets the object state to the original state when loaded from or
'                   last saved to the repository
' Archive()         Replaces the current value of the object to the one in collection
' CheckDirty()      Checks if the values are different in the collection and the current object
' Init()            Initializes the member variables to their default values
'
'-------------------------------------------------------------------------------

Namespace MUSTER.Info
    <Serializable()> _
    Public Class PipeInfo
#Region "Private member variables"
        Private nPipeId As Integer
        Private nPipeIndex As Integer
        Private nFacilityID As Integer
        Private nTankID As Integer
        Private strALLDTest As String
        Private dtALLDTestDate As Date
        Private nCASNumber As Integer
        Private nClosureStatusDesc As Integer
        Private nClosureType As Integer
        Private nCompPrimary As Integer
        Private nCompSecondary As Integer
        Private boolContainSumpDisp As Boolean
        Private boolContainSumpTank As Boolean
        Private dtDateClosed As Date
        Private dtDateLastUsed As Date
        Private dtDateClosureRecd As Date
        Private dtDateRecd As Date
        Private dtDateSigned As Date
        Private nALLDType As Integer
        Private nInertMaterial As Integer
        Private dtLCPInstallDate As Date
        Private nLicenseeID As Integer
        Private nContractorID As Integer
        Private dtLTTDate As Date
        Private dtPipeCPTest As Date
        Private dtShearTest As Date
        Private dtPipeSecInsp As Date
        Private dtPipeElecInsp As Date
        Private nPipeCPType As Integer
        Private dtPipeInstallDate As Date
        Private nPipeLD As Integer
        Private nPipeManufacturer As Integer
        Private nPipeMatDesc As Integer
        Private nPipeModDesc As Integer
        Private strPipeOtherMaterial As String
        Private nPipeStatusDesc As Integer
        Private nPipeTypeDesc As Integer
        Private strPipingComments As String
        Private dtPipeInstallationPlannedFor As Date
        Private dtPlacedInServiceDate As Date
        Private nSubstanceComments As Integer
        Private nSubstanceDesc As Integer
        Private dtTermCPLastTested As Date
        Private nTermCPTypeTank As Integer
        Private nTermCPTypeDisp As Integer
        Private dtPipeCPInstalledDate As Date
        Private dtTermCPInstalledDate As Date
        Private nTermTypeDisp As Integer
        Private nTermTypeTank As Integer
        Private nParentPipeID As Integer
        Private nHasExtensions As Boolean
        Private boolDeleted As Boolean
        Private strCreatedBy As String = String.Empty
        Private dtCreatedOn As Date = DateTime.Now.ToShortDateString
        Private strModifiedBy As String = String.Empty
        Private dtModifiedOn As Date = DateTime.Now.ToShortDateString
        Private nCompartmentNumber As Integer
        Private nCompartmentSubstance As Integer
        Private nCompartmentCERCLA As Integer
        Private nCompartmentFuelType As Integer
        Private nTankSiteID As Integer
        Private nAttachedPipeId As Integer
        Private bolPOU As Boolean
        Private bolNonPre88 As Boolean
        Private bolFacPowerOff As Boolean

        Private onPipeId As Integer
        Private onPipeIndex As Integer
        Private onFacilityID As Integer
        Private onTankID As Integer
        Private ostrALLDTest As String
        Private odtALLDTestDate As Date
        Private onCASNumber As Integer
        Private onClosureStatusDesc As Integer
        Private onClosureType As Integer
        Private onCompPrimary As Integer
        Private onCompSecondary As Integer
        Private oboolContainSumpDisp As Boolean
        Private oboolContainSumpTank As Boolean
        Private odtDateClosed As Date
        Private odtDateLastUsed As Date
        Private odtDateClosureRecd As Date
        Private odtDateRecd As Date
        Private odtDateSigned As Date
        Private onALLDType As Integer
        Private onInertMaterial As Integer
        Private odtLCPInstallDate As Date
        Private onLicenseeID As Integer
        Private onContractorID As Integer
        Private odtLTTDate As Date
        Private odtPipeCPTest As Date
        Private odtShearTest As Date
        Private odtPipeSecInsp As Date
        Private odtPipeElecInsp As Date
        Private onPipeCPType As Integer
        Private odtPipeInstallDate As Date
        Private onPipeLD As Integer
        Private onPipeManufacturer As Integer
        Private onPipeMatDesc As Integer
        Private onPipeModDesc As Integer
        Private ostrPipeOtherMaterial As String
        Private onPipeStatusDesc As Integer
        Private onPipeTypeDesc As Integer
        Private ostrPipingComments As String
        Private odtPipeInstallationPlannedFor As Date
        Private odtPlacedInServiceDate As Date
        Private onSubstanceComments As Integer
        Private onSubstanceDesc As Integer
        Private odtTermCPLastTested As Date
        Private onTermCPTypeTank As Integer
        Private onTermCPTypeDisp As Integer
        Private odtPipeCPInstalledDate As Date
        Private odtTermCPInstalledDate As Date
        Private onTermTypeDisp As Integer
        Private onTermTypeTank As Integer
        Private oboolDeleted As Boolean
        Private ostrCreatedBy As String
        Private odtCreatedOn As Date
        Private ostrModifiedBy As String
        Private odtModifiedOn As Date
        Private onCompartmentNumber As Integer
        Private onCompartmentSubstance As Integer
        Private onCompartmentCERCLA As Integer
        Private onCompartmentFuelType As Integer
        Private onTankSiteID As Integer
        Private onAttachedPipeId As Integer
        Private onParentPipeID As Integer
        Private onHasExtensions As Boolean



        Private nFacCapStatus As Integer = 0
        Private dtDataAge As DateTime
        Private nAgeThreshold As Int16 = 5
        Private bolShowDeleted As Boolean = False
        Private bolIsDirty As Boolean = False
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
        'added by kiran
        Dim colComments As MUSTER.Info.CommentsCollection
        'end changes
#End Region
#Region "Public Events"
        Public Event evtPipeInfoChanged(ByVal DirtyState As Boolean)
        'Public Event InfoBecameDirty(ByVal DirtyState As Boolean, ByVal strSrc As String)
#End Region
#Region "Constructors"
        Sub New()
            MyBase.new()
            dtDataAge = Now()
            Me.Init()
            'added by kiran
            colComments = New MUSTER.Info.CommentsCollection
            'end changes
        End Sub
        Sub New(ByVal pipeID As Integer, _
                ByVal pipeIndex As Integer, _
                ByVal facilityID As Integer, _
                ByVal tankID As Integer, _
                ByVal alldTest As String, _
                ByVal alldTestDate As Date, _
                ByVal casNumber As Integer, _
                ByVal closureStatusDesc As Integer, _
                ByVal closureType As Integer, _
                ByVal compPrimary As Integer, _
                ByVal compSecondary As Integer, _
                ByVal containSumpDisp As Boolean, _
                ByVal containSumpTank As Boolean, _
                ByVal dateClosed As Date, _
                ByVal dateLastUsed As Date, _
                ByVal dateClosureRecd As Date, _
                ByVal dateRecd As Date, _
                ByVal dateSigned As Date, _
                ByVal alldType As Integer, _
                ByVal inertMaterial As Integer, _
                ByVal lcpInstallDate As Date, _
                ByVal licenseeID As Integer, _
                ByVal contractorID As Integer, _
                ByVal lttDate As Date, _
                ByVal pipeCPTest As Date, _
                ByVal shearTest As Date, _
                ByVal pipeSecInsp As Date, _
                ByVal pipeElecInsp As Date, _
                ByVal pipeCPType As Integer, _
                ByVal pipeInstallDate As Date, _
                ByVal pipeLD As Integer, _
                ByVal pipeManufacturer As Integer, _
                ByVal pipeMatDesc As Integer, _
                ByVal pipeModDesc As Integer, _
                ByVal pipeOtherMaterial As String, _
                ByVal pipeStatusDesc As Integer, _
                ByVal pipeTypeDesc As Integer, _
                ByVal pipingComments As String, _
                ByVal pipeInstallationsPlannedFor As Date, _
                ByVal placedInServiceDate As Date, _
                ByVal substanceComments As Integer, _
                ByVal substanceDesc As Integer, _
                ByVal termCPLastTested As Date, _
                ByVal termCPTypeTank As Integer, _
                ByVal termCPTypeDisp As Integer, _
                ByVal pipeCPInstalledDate As Date, _
                ByVal termCPInstalledDate As Date, _
                ByVal termTypeDisp As Integer, _
                ByVal termTypeTank As Integer, _
                ByVal deleted As Boolean, _
                ByVal createdBy As String, _
                ByVal createdOn As Date, _
                ByVal modifiedBy As String, _
                ByVal modifiedOn As Date, _
                Optional ByVal compNum As Integer = 0, _
                Optional ByVal ParentPipeId As Integer = 0, _
                Optional ByVal HasExtensions As Boolean = False)
            onPipeId = pipeID
            onPipeIndex = pipeIndex
            onFacilityID = facilityID
            onTankID = tankID
            ostrALLDTest = alldTest
            odtALLDTestDate = alldTestDate
            onCASNumber = casNumber
            onClosureStatusDesc = closureStatusDesc
            onClosureType = closureType
            onCompPrimary = compPrimary
            onCompSecondary = compSecondary
            oboolContainSumpDisp = containSumpDisp
            oboolContainSumpTank = containSumpTank
            odtDateClosed = dateClosed.Date
            odtDateLastUsed = dateLastUsed.Date
            odtDateClosureRecd = dateClosureRecd.Date
            odtDateRecd = dateRecd.Date
            odtDateSigned = dateSigned.Date
            onALLDType = alldType
            onInertMaterial = inertMaterial
            odtLCPInstallDate = lcpInstallDate.Date
            onLicenseeID = licenseeID
            onContractorID = contractorID
            odtLTTDate = lttDate.Date
            odtPipeCPTest = pipeCPTest.Date
            odtShearTest = shearTest.Date
            odtPipeSecInsp = pipeSecInsp.Date
            odtPipeElecInsp = pipeElecInsp.Date
            onPipeCPType = PipeCPType
            odtPipeInstallDate = PipeInstallDate.Date
            onPipeLD = PipeLD
            onPipeManufacturer = PipeManufacturer
            onPipeMatDesc = PipeMatDesc
            onPipeModDesc = PipeModDesc
            ostrPipeOtherMaterial = PipeOtherMaterial
            onPipeStatusDesc = PipeStatusDesc
            onPipeTypeDesc = PipeTypeDesc
            ostrPipingComments = PipingComments
            odtPipeInstallationPlannedFor = pipeInstallationsPlannedFor.Date
            odtPlacedInServiceDate = PlacedInServiceDate.Date
            onSubstanceComments = SubstanceComments
            onSubstanceDesc = SubstanceDesc
            odtTermCPLastTested = TermCPLastTested.Date
            onTermCPTypeTank = TermCPTypeTank
            onTermCPTypeDisp = TermCPTypeDisp
            odtPipeCPInstalledDate = PipeCPInstalledDate.Date
            odtTermCPInstalledDate = TermCPInstalledDate.Date
            onTermTypeDisp = TermTypeDisp
            onTermTypeTank = TermTypeTank
            oboolDeleted = Deleted
            ostrCreatedBy = CreatedBy
            odtCreatedOn = CreatedOn
            ostrModifiedBy = ModifiedBy
            odtModifiedOn = ModifiedOn
            onCompartmentNumber = 0
            onCompartmentSubstance = 0
            onCompartmentCERCLA = 0
            onCompartmentFuelType = 0
            onTankSiteID = 0
            onCompartmentNumber = compNum
            onAttachedPipeId = AttachedPipeID
            onParentPipeID = ParentPipeID
            onHasExtensions = HasExtensions
            dtDataAge = Now()
            'added by kiran
            colComments = New MUSTER.Info.CommentsCollection
            'end changes
            Me.Reset()
        End Sub
        Sub New(ByVal drPipe As DataRow)
            Try
                'IIf(drPipe.Item("") Is System.DBNull.Value, False, drPipe.Item(""))
                onPipeId = drPipe.Item("PIPE_ID")
                onPipeIndex = drPipe.Item("PIPE_INDEX")
                onFacilityID = drPipe.Item("FACILITY_ID")
                onTankID = IIf(drPipe.Item("COMPARTMENTS_PIPES_TANKID") Is System.DBNull.Value, 0, drPipe.Item("COMPARTMENTS_PIPES_TANKID"))
                ostrALLDTest = IIf(drPipe.Item("ALLD_TEST") Is System.DBNull.Value, String.Empty, drPipe.Item("ALLD_TEST"))
                odtALLDTestDate = IIf(drPipe.Item("ALLD_TEST_DATE") Is System.DBNull.Value, CDate("01/01/0001"), drPipe.Item("ALLD_TEST_DATE"))
                odtALLDTestDate = odtALLDTestDate.Date
                onCASNumber = IIf(drPipe.Item("CAS_NUMBER") Is System.DBNull.Value, 0, drPipe.Item("CAS_NUMBER"))
                onClosureStatusDesc = IIf(drPipe.Item("CLOSURE_STATUS_DESC") Is DBNull.Value, 0, drPipe.Item("CLOSURE_STATUS_DESC"))
                onClosureType = IIf(drPipe.Item("CLOSURETYPE") Is DBNull.Value, 0, drPipe.Item("CLOSURETYPE"))
                onCompPrimary = IIf(drPipe.Item("COMPOSITE_PRIMARY") Is DBNull.Value, 0, drPipe.Item("COMPOSITE_PRIMARY"))
                onCompSecondary = IIf(drPipe.Item("COMPOSITE_SECONDARY") Is DBNull.Value, 0, drPipe.Item("COMPOSITE_SECONDARY"))
                oboolContainSumpDisp = IIf(drPipe.Item("CONTAIN_SUMPDISP") Is DBNull.Value, False, drPipe.Item("CONTAIN_SUMPDISP"))
                oboolContainSumpTank = IIf(drPipe.Item("CONTAIN_SUMPTANK") Is DBNull.Value, False, drPipe.Item("CONTAIN_SUMPTANK"))
                odtDateClosed = IIf(drPipe.Item("DATE_CLOSED") Is System.DBNull.Value, CDate("01/01/0001"), drPipe.Item("DATE_CLOSED"))
                odtDateClosed = odtDateClosed.Date
                odtDateLastUsed = IIf(drPipe.Item("DATE_LAST_USED") Is System.DBNull.Value, CDate("01/01/0001"), drPipe.Item("DATE_LAST_USED"))
                odtDateLastUsed = odtDateLastUsed.Date
                odtDateClosureRecd = IIf(drPipe.Item("DATE_CLOSURE_RECD") Is System.DBNull.Value, CDate("01/01/0001"), drPipe.Item("DATE_CLOSURE_RECD"))
                odtDateClosureRecd = odtDateClosureRecd.Date
                odtDateRecd = IIf(drPipe.Item("DATE_RECD") Is System.DBNull.Value, CDate("01/01/0001"), drPipe.Item("DATE_RECD"))
                odtDateRecd = odtDateRecd.Date
                odtDateSigned = IIf(drPipe.Item("DATE_SIGNED") Is System.DBNull.Value, CDate("01/01/0001"), drPipe.Item("DATE_SIGNED"))
                odtDateSigned = odtDateSigned.Date
                onALLDType = IIf(drPipe.Item("ALLD_TYPE") Is DBNull.Value, 0, drPipe.Item("ALLD_TYPE"))
                onInertMaterial = IIf(drPipe.Item("INERT_MATERIAL") Is DBNull.Value, 0, drPipe.Item("INERT_MATERIAL"))
                odtLCPInstallDate = IIf(drPipe.Item("LCP_INSTALL_DATE") Is System.DBNull.Value, CDate("01/01/0001"), drPipe.Item("LCP_INSTALL_DATE"))
                odtLCPInstallDate = odtLCPInstallDate.Date
                onLicenseeID = IIf(drPipe.Item("LICENSEE_ID") Is System.DBNull.Value, 0, drPipe.Item("LICENSEE_ID"))
                onContractorID = IIf(drPipe.Item("CONTRACTOR_ID") Is System.DBNull.Value, 0, drPipe.Item("CONTRACTOR_ID"))
                odtLTTDate = IIf(drPipe.Item("LTT_DATE") Is System.DBNull.Value, CDate("01/01/0001"), drPipe.Item("LTT_DATE"))
                odtLTTDate = odtLTTDate.Date
                odtPipeCPTest = IIf(drPipe.Item("PIPE_CP_TEST") Is System.DBNull.Value, CDate("01/01/0001"), drPipe.Item("PIPE_CP_TEST"))
                odtPipeCPTest = odtPipeCPTest.Date
                odtShearTest = IIf(drPipe.Item("DateSheerValueTest") Is System.DBNull.Value, CDate("01/01/0001"), drPipe.Item("DateSheerValueTest"))
                odtShearTest = odtShearTest.Date
                odtPipeSecInsp = IIf(drPipe.Item("DateSecondaryContainmentInspect") Is System.DBNull.Value, CDate("01/01/0001"), drPipe.Item("DateSecondaryContainmentInspect"))
                odtPipeSecInsp = odtPipeSecInsp.Date
                odtPipeElecInsp = IIf(drPipe.Item("DateElectronicDeviceInspect") Is System.DBNull.Value, CDate("01/01/0001"), drPipe.Item("DateElectronicDeviceInspect"))
                odtPipeElecInsp = odtPipeElecInsp.Date
                onPipeCPType = IIf(drPipe.Item("PIPE_CP_TYPE") Is DBNull.Value, 0, drPipe.Item("PIPE_CP_TYPE"))
                odtPipeInstallDate = IIf(drPipe.Item("PIPE_INSTALL_DATE") Is System.DBNull.Value, CDate("01/01/0001"), drPipe.Item("PIPE_INSTALL_DATE"))
                odtPipeInstallDate = odtPipeInstallDate.Date
                onPipeLD = IIf(drPipe.Item("PIPE_LD") Is DBNull.Value, 0, drPipe.Item("PIPE_LD"))
                onPipeManufacturer = IIf(drPipe.Item("PIPE_MANUFACTURER") Is DBNull.Value, 0, drPipe.Item("PIPE_MANUFACTURER"))
                onPipeMatDesc = IIf(drPipe.Item("PIPE_MAT_DESC") Is DBNull.Value, 0, drPipe.Item("PIPE_MAT_DESC"))
                onPipeModDesc = drPipe.Item("PIPE_MOD_DESC")
                ostrPipeOtherMaterial = IIf(drPipe.Item("PIPE_OTHER_MATERIAL") Is DBNull.Value, String.Empty, drPipe.Item("PIPE_OTHER_MATERIAL"))
                onPipeStatusDesc = IIf(drPipe.Item("PIPE_STATUS_DESC") Is DBNull.Value, 0, drPipe.Item("PIPE_STATUS_DESC"))
                onPipeTypeDesc = IIf(drPipe.Item("PIPE_TYPE_DESC") Is DBNull.Value, 0, drPipe.Item("PIPE_TYPE_DESC"))
                ostrPipingComments = IIf(drPipe.Item("PIPING_COMMENTS") Is DBNull.Value, String.Empty, drPipe.Item("PIPING_COMMENTS"))
                odtPipeInstallationPlannedFor = IIf(drPipe.Item("PIPE_INSTALLATION_PLANNED_FOR") Is System.DBNull.Value, CDate("01/01/0001"), drPipe.Item("PIPE_INSTALLATION_PLANNED_FOR"))
                odtPipeInstallationPlannedFor = odtPipeInstallationPlannedFor.Date
                odtPlacedInServiceDate = IIf(drPipe.Item("PLACED_IN_SERVICE_DATE") Is System.DBNull.Value, CDate("01/01/0001"), drPipe.Item("PLACED_IN_SERVICE_DATE"))
                odtPlacedInServiceDate = odtPlacedInServiceDate.Date
                onSubstanceComments = IIf(drPipe.Item("SUBSTANCE_COMMENTS") Is DBNull.Value, 0, drPipe.Item("SUBSTANCE_COMMENTS"))
                onSubstanceDesc = IIf(drPipe.Item("SUBSTANCE_DESC") Is DBNull.Value, 0, drPipe.Item("SUBSTANCE_DESC"))
                odtTermCPLastTested = IIf(drPipe.Item("TERM_CP_LAST_TESTED") Is System.DBNull.Value, CDate("01/01/0001"), drPipe.Item("TERM_CP_LAST_TESTED"))
                odtTermCPLastTested = odtTermCPLastTested.Date
                onTermCPTypeTank = IIf(drPipe.Item("TERM_CP_TYPE_TANK") Is DBNull.Value, 0, drPipe.Item("TERM_CP_TYPE_TANK"))
                onTermCPTypeDisp = IIf(drPipe.Item("TERM_CP_TYPE_DISP") Is DBNull.Value, 0, drPipe.Item("TERM_CP_TYPE_DISP"))
                odtPipeCPInstalledDate = IIf(drPipe.Item("PIPE_CP_INSTALLED_DATE") Is System.DBNull.Value, CDate("01/01/0001"), drPipe.Item("PIPE_CP_INSTALLED_DATE"))
                odtPipeCPInstalledDate = odtPipeCPInstalledDate.Date
                odtTermCPInstalledDate = IIf(drPipe.Item("TERMINATION_CP_INSTALLED_DATE") Is System.DBNull.Value, CDate("01/01/0001"), drPipe.Item("TERMINATION_CP_INSTALLED_DATE"))
                odtTermCPInstalledDate = odtTermCPInstalledDate.Date
                onTermTypeDisp = IIf(drPipe.Item("TERMINATION_TYPE_DISP") Is DBNull.Value, 0, drPipe.Item("TERMINATION_TYPE_DISP"))
                onTermTypeTank = IIf(drPipe.Item("TERMINATION_TYPE_TANK") Is DBNull.Value, 0, drPipe.Item("TERMINATION_TYPE_TANK"))
                oboolDeleted = IIf(drPipe.Item("DELETED") Is DBNull.Value, False, drPipe.Item("DELETED"))
                ostrCreatedBy = IIf(drPipe.Item("CREATED_BY") Is System.DBNull.Value, False, drPipe.Item("CREATED_BY"))
                odtCreatedOn = IIf(drPipe.Item("DATE_CREATED") Is System.DBNull.Value, CDate("01/01/0001"), drPipe.Item("DATE_CREATED"))
                ostrModifiedBy = IIf(drPipe.Item("LAST_EDITED_BY") Is System.DBNull.Value, False, drPipe.Item("LAST_EDITED_BY"))
                odtModifiedOn = IIf(drPipe.Item("DATE_LAST_EDITED") Is System.DBNull.Value, CDate("01/01/0001"), drPipe.Item("DATE_LAST_EDITED"))
                onCompartmentNumber = IIf(drPipe.Item("COMPARTMENT_NUMBER") Is System.DBNull.Value, 0, drPipe.Item("COMPARTMENT_NUMBER"))
                onCompartmentSubstance = IIf(drPipe.Item("SUBSTANCE") Is System.DBNull.Value, 0, drPipe.Item("SUBSTANCE"))
                onCompartmentCERCLA = IIf(drPipe.Item("CERCLA#") Is System.DBNull.Value, 0, drPipe.Item("CERCLA#"))
                onCompartmentFuelType = IIf(drPipe.Item("FUEL_TYPE_ID") Is System.DBNull.Value, 0, drPipe.Item("FUEL_TYPE_ID"))
                onTankSiteID = 0
                onAttachedPipeId = 0
                onParentPipeID = IIf(drPipe.Item("PARENT_PIPE_ID") Is System.DBNull.Value, 0, drPipe.Item("PARENT_PIPE_ID"))
                onHasExtensions = IIf(drPipe.Item("HAS_EXTENSIONS") Is System.DBNull.Value, False, (drPipe.Item("HAS_EXTENSIONS") = 1))
                dtDataAge = Now()
                'added by kiran
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
            If nPipeId >= 0 Then
                nPipeId = onPipeId
            End If
            nPipeIndex = onPipeIndex
            nFacilityID = onFacilityID
            nTankID = onTankID
            strALLDTest = ostrALLDTest
            dtALLDTestDate = odtALLDTestDate
            nCASNumber = onCASNumber
            nClosureStatusDesc = onClosureStatusDesc
            nClosureType = onClosureType
            nCompPrimary = onCompPrimary
            nCompSecondary = onCompSecondary
            boolContainSumpDisp = oboolContainSumpDisp
            boolContainSumpTank = oboolContainSumpTank
            dtDateClosed = odtDateClosed
            dtDateLastUsed = odtDateLastUsed
            dtDateClosureRecd = odtDateClosureRecd
            dtDateRecd = odtDateRecd
            dtDateSigned = odtDateSigned
            nALLDType = onALLDType
            nInertMaterial = onInertMaterial
            dtLCPInstallDate = odtLCPInstallDate
            nLicenseeID = onLicenseeID
            nContractorID = onContractorID
            dtLTTDate = odtLTTDate
            dtPipeCPTest = odtPipeCPTest
            dtShearTest = odtShearTest
            dtPipeSecInsp = odtPipeSecInsp
            dtpipeElecInsp = odtPipeElecInsp
            nPipeCPType = onPipeCPType
            dtPipeInstallDate = odtPipeInstallDate
            nPipeLD = onPipeLD
            nPipeManufacturer = onPipeManufacturer
            nPipeMatDesc = onPipeMatDesc
            nPipeModDesc = onPipeModDesc
            strPipeOtherMaterial = ostrPipeOtherMaterial
            nPipeStatusDesc = onPipeStatusDesc
            nPipeTypeDesc = onPipeTypeDesc
            strPipingComments = ostrPipingComments
            dtPipeInstallationPlannedFor = odtPipeInstallationPlannedFor
            dtPlacedInServiceDate = odtPlacedInServiceDate
            nSubstanceComments = onSubstanceComments
            nSubstanceDesc = onSubstanceDesc
            dtTermCPLastTested = odtTermCPLastTested
            nTermCPTypeTank = onTermCPTypeTank
            nTermCPTypeDisp = onTermCPTypeDisp
            dtPipeCPInstalledDate = odtPipeCPInstalledDate
            dtTermCPInstalledDate = odtTermCPInstalledDate
            nTermTypeDisp = onTermTypeDisp
            nTermTypeTank = onTermTypeTank
            boolDeleted = oboolDeleted
            strCreatedBy = ostrCreatedBy
            dtCreatedOn = odtCreatedOn
            strModifiedBy = ostrModifiedBy
            dtModifiedOn = odtModifiedOn
            nCompartmentNumber = onCompartmentNumber
            nCompartmentSubstance = onCompartmentSubstance
            nCompartmentCERCLA = onCompartmentCERCLA
            nCompartmentFuelType = onCompartmentFuelType
            nTankSiteID = onTankSiteID
            nAttachedPipeId = onAttachedPipeId
            nParentPipeID = onParentPipeID
            nHasExtensions = onHasExtensions

            bolIsDirty = False
            RaiseEvent evtPipeInfoChanged(bolIsDirty)
        End Sub
        Public Sub Archive()
            onPipeId = nPipeId
            onPipeIndex = nPipeIndex
            onFacilityID = nFacilityID
            onTankID = nTankID
            ostrALLDTest = strALLDTest
            odtALLDTestDate = dtALLDTestDate
            onCASNumber = nCASNumber
            onClosureStatusDesc = nClosureStatusDesc
            onClosureType = nClosureType
            onCompPrimary = nCompPrimary
            onCompSecondary = nCompSecondary
            oboolContainSumpDisp = boolContainSumpDisp
            oboolContainSumpTank = boolContainSumpTank
            odtDateClosed = dtDateClosed
            odtDateLastUsed = dtDateLastUsed
            odtDateClosureRecd = dtDateClosureRecd
            odtDateRecd = dtDateRecd
            odtDateSigned = dtDateSigned
            onALLDType = nALLDType
            onInertMaterial = nInertMaterial
            odtLCPInstallDate = dtLCPInstallDate
            onLicenseeID = nLicenseeID
            onContractorID = nContractorID
            odtLTTDate = dtLTTDate
            odtPipeCPTest = dtPipeCPTest
            odtShearTest = dtShearTest
            odtPipeSecInsp = dtPipeSecInsp
            odtPipeElecInsp = dtpipeElecInsp
            onPipeCPType = nPipeCPType
            odtPipeInstallDate = dtPipeInstallDate
            onPipeLD = nPipeLD
            onPipeManufacturer = nPipeManufacturer
            onPipeMatDesc = nPipeMatDesc
            onPipeModDesc = nPipeModDesc
            ostrPipeOtherMaterial = strPipeOtherMaterial
            onPipeStatusDesc = nPipeStatusDesc
            onPipeTypeDesc = nPipeTypeDesc
            ostrPipingComments = strPipingComments
            odtPipeInstallationPlannedFor = dtPipeInstallationPlannedFor
            odtPlacedInServiceDate = dtPlacedInServiceDate
            onSubstanceComments = nSubstanceComments
            onSubstanceDesc = nSubstanceDesc
            odtTermCPLastTested = dtTermCPLastTested
            onTermCPTypeTank = nTermCPTypeTank
            onTermCPTypeDisp = nTermCPTypeDisp
            odtPipeCPInstalledDate = dtPipeCPInstalledDate
            odtTermCPInstalledDate = dtTermCPInstalledDate
            onTermTypeDisp = nTermTypeDisp
            onTermTypeTank = nTermTypeTank
            oboolDeleted = boolDeleted
            ostrCreatedBy = strCreatedBy
            odtCreatedOn = dtCreatedOn
            ostrModifiedBy = strModifiedBy
            odtModifiedOn = dtModifiedOn
            onCompartmentNumber = nCompartmentNumber
            onCompartmentSubstance = nCompartmentSubstance
            onCompartmentCERCLA = nCompartmentCERCLA
            onCompartmentFuelType = nCompartmentFuelType
            onTankSiteID = nTankSiteID
            onAttachedPipeId = nAttachedPipeId
            onParentPipeID = nParentPipeID
            onHasExtensions = nHasExtensions
            bolIsDirty = False
        End Sub
#End Region
#Region "Private Operations"
        Private Sub CheckDirty()
            Dim obolIsDirty As Boolean = bolIsDirty
            '(nFacilityID <> onFacilityID) Or _
            '(nTankID <> onTankID) Or _
            '(nSubstanceDesc <> onSubstanceDesc) Or _
            bolIsDirty = (strALLDTest <> ostrALLDTest) Or _
                            (dtALLDTestDate <> odtALLDTestDate) Or _
                            (nCASNumber <> onCASNumber) Or _
                            (nClosureStatusDesc <> onClosureStatusDesc) Or _
                            (nClosureType <> onClosureType) Or _
                            (nCompPrimary <> onCompPrimary) Or _
                            (nCompSecondary <> onCompSecondary) Or _
                            (boolContainSumpDisp <> oboolContainSumpDisp) Or _
                            (boolContainSumpTank <> oboolContainSumpTank) Or _
                            (dtDateClosed <> odtDateClosed) Or _
                            (dtDateLastUsed <> odtDateLastUsed) Or _
                            (dtDateClosureRecd <> odtDateClosureRecd) Or _
                            (dtDateRecd <> odtDateRecd) Or _
                            (dtDateSigned <> odtDateSigned) Or _
                            (nALLDType <> onALLDType) Or _
                            (nInertMaterial <> onInertMaterial) Or _
                            (dtLCPInstallDate <> odtLCPInstallDate) Or _
                            (nLicenseeID <> onLicenseeID) Or _
                            (nContractorID <> onContractorID) Or _
                            (dtLTTDate <> odtLTTDate) Or _
                            (dtPipeCPTest <> odtPipeCPTest) Or _
                            (dtShearTest <> odtShearTest) Or _
                            (dtPipeSecInsp <> odtPipeSecInsp) Or _
                            (dtPipeElecInsp <> odtPipeElecInsp) Or _
                            (nPipeCPType <> onPipeCPType) Or _
                            (dtPipeInstallDate <> odtPipeInstallDate) Or _
                            (nPipeLD <> onPipeLD) Or _
                            (nPipeManufacturer <> onPipeManufacturer) Or _
                            (nPipeMatDesc <> onPipeMatDesc) Or _
                            (nPipeModDesc <> onPipeModDesc) Or _
                            (strPipeOtherMaterial <> ostrPipeOtherMaterial) Or _
                            (nPipeStatusDesc <> onPipeStatusDesc) Or _
                            (nPipeTypeDesc <> onPipeTypeDesc) Or _
                            (strPipingComments <> ostrPipingComments) Or _
                            (dtPipeInstallationPlannedFor <> odtPipeInstallationPlannedFor) Or _
                            (dtPlacedInServiceDate <> odtPlacedInServiceDate) Or _
                            (nSubstanceComments <> onSubstanceComments) Or _
                            (dtTermCPLastTested <> odtTermCPLastTested) Or _
                            (nTermCPTypeTank <> onTermCPTypeTank) Or _
                            (nTermCPTypeDisp <> onTermCPTypeDisp) Or _
                            (dtPipeCPInstalledDate <> odtPipeCPInstalledDate) Or _
                            (dtTermCPInstalledDate <> odtTermCPInstalledDate) Or _
                            (nTermTypeDisp <> onTermTypeDisp) Or _
                            (nTermTypeTank <> onTermTypeTank) Or _
                            (nParentPipeID <> onParentPipeID) Or _
                            (nHasExtensions <> onHasExtensions) Or _
                            (boolDeleted <> oboolDeleted)
            If obolIsDirty <> bolIsDirty Then
                'MsgBox("Info F:" + FacilityID.ToString + " TI:" + TankSiteID.ToString + " PID:" + ID)
                RaiseEvent evtPipeInfoChanged(bolIsDirty)
            End If
        End Sub
        Private Sub Init()
            onPipeId = 0
            onPipeIndex = 0
            onFacilityID = 0
            onTankID = 0
            ostrALLDTest = String.Empty
            odtALLDTestDate = CDate("01/01/0001")
            onCASNumber = 0
            onClosureStatusDesc = 0
            onClosureType = 0
            onCompPrimary = 0
            onCompSecondary = 0
            oboolContainSumpDisp = False
            oboolContainSumpTank = False
            odtDateClosed = CDate("01/01/0001")
            odtDateLastUsed = CDate("01/01/0001")
            odtDateClosureRecd = CDate("01/01/0001")
            odtDateRecd = CDate("01/01/0001")
            odtDateSigned = CDate("01/01/0001")
            onALLDType = 0
            onInertMaterial = 0
            odtLCPInstallDate = CDate("01/01/0001")
            onLicenseeID = 0
            onContractorID = 0
            odtLTTDate = CDate("01/01/0001")
            odtPipeCPTest = CDate("01/01/0001")
            odtShearTest = CDate("01/01/0001")
            odtPipeSecInsp = CDate("01/01/0001")
            odtPipeElecInsp = CDate("01/01/0001")
            onPipeCPType = 0
            odtPipeInstallDate = CDate("01/01/0001")
            onPipeLD = 0
            onPipeManufacturer = 0
            onPipeMatDesc = 0
            onPipeModDesc = 0
            ostrPipeOtherMaterial = String.Empty
            onPipeStatusDesc = 0
            onPipeTypeDesc = 0
            ostrPipingComments = String.Empty
            odtPipeInstallationPlannedFor = CDate("01/01/0001")
            odtPlacedInServiceDate = CDate("01/01/0001")
            onSubstanceComments = 0
            onSubstanceDesc = 0
            odtTermCPLastTested = CDate("01/01/0001")
            onTermCPTypeTank = 0
            onTermCPTypeDisp = 0
            odtPipeCPInstalledDate = CDate("01/01/0001")
            odtTermCPInstalledDate = CDate("01/01/0001")
            onTermTypeDisp = 0
            onTermTypeTank = 0
            oboolDeleted = False
            onCompartmentNumber = 0
            onCompartmentSubstance = 0
            onCompartmentCERCLA = 0
            onCompartmentFuelType = 0
            onTankSiteID = 0
            onAttachedPipeId = 0
            onParentPipeID = 0
            onHasExtensions = 0
            bolPOU = False
            bolNonPre88 = False
            bolFacPowerOff = False
            odtCreatedOn = CDate("01/01/0001")
            odtModifiedOn = CDate("01/01/0001")
            nFacCapStatus = 0
            Me.Reset()
        End Sub
#End Region
#Region "Exposed Attributes"
        'added by kiran
        Public Property commentsCollection() As MUSTER.Info.CommentsCollection
            Get
                Return colComments
            End Get
            Set(ByVal Value As MUSTER.Info.CommentsCollection)
                colComments = Value
            End Set
        End Property
        'end changes
        Public Property ID() As String
            Get
                Return TankID.ToString + "|" + CompartmentNumber.ToString + "|" + PipeID.ToString
            End Get
            Set(ByVal Value As String)
                Dim arrVals() As String
                arrVals = Value.Split("|")
                nTankID = Integer.Parse(arrVals(0))
                nCompartmentNumber = Integer.Parse(arrVals(1))
                nPipeId = Integer.Parse(arrVals(2))
                Me.CheckDirty()
            End Set
        End Property
        Public Property PipeID() As Integer
            Get
                Return nPipeId
            End Get
            Set(ByVal Value As Integer)
                nPipeId = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property ParentPipeID() As Integer
            Get
                Return nParentPipeID
            End Get
            Set(ByVal Value As Integer)
                nParentPipeID = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property HasExtensions() As Boolean
            Get
                Return nHasExtensions
            End Get
            Set(ByVal Value As Boolean)
                nHasExtensions = Value
            End Set
        End Property

        Public ReadOnly Property HasParent() As Boolean

            Get
                Return (ParentPipeID > 0)
            End Get

        End Property

        Public Property Index() As Integer
            Get
                Return nPipeIndex
            End Get
            Set(ByVal Value As Integer)
                nPipeIndex = Value
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
        Public Property TankID() As Integer
            Get
                Return nTankID
            End Get
            Set(ByVal Value As Integer)
                nTankID = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property ALLDTest() As String
            Get
                Return strALLDTest
            End Get
            Set(ByVal Value As String)
                strALLDTest = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property ALLDTestDate() As Date
            Get
                Return dtALLDTestDate.Date
            End Get
            Set(ByVal Value As Date)
                dtALLDTestDate = Value.Date
                Me.CheckDirty()
            End Set
        End Property
        Public Property CASNumber() As Integer
            Get
                Return nCASNumber
            End Get
            Set(ByVal Value As Integer)
                nCASNumber = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property ClosureStatusDesc() As Integer
            Get
                Return nClosureStatusDesc
            End Get
            Set(ByVal Value As Integer)
                nClosureStatusDesc = Value
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
        Public Property CompPrimary() As Integer
            Get
                Return nCompPrimary
            End Get
            Set(ByVal Value As Integer)
                nCompPrimary = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property CompSecondary() As Integer
            Get
                Return nCompSecondary
            End Get
            Set(ByVal Value As Integer)
                nCompSecondary = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property ContainSumpDisp() As Boolean
            Get
                Return boolContainSumpDisp
            End Get
            Set(ByVal Value As Boolean)
                boolContainSumpDisp = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property ContainSumpTank() As Boolean
            Get
                Return boolContainSumpTank
            End Get
            Set(ByVal Value As Boolean)
                boolContainSumpTank = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property DateClosed() As Date
            Get
                Return dtDateClosed.Date
            End Get
            Set(ByVal Value As Date)
                dtDateClosed = Value.Date
                Me.CheckDirty()
            End Set
        End Property
        Public Property DateLastUsed() As Date
            Get
                Return dtDateLastUsed.Date
            End Get
            Set(ByVal Value As Date)
                dtDateLastUsed = Value.Date
                Me.CheckDirty()
            End Set
        End Property
        Public Property DateClosureRecd() As Date
            Get
                Return dtDateClosureRecd.Date
            End Get
            Set(ByVal Value As Date)
                dtDateClosureRecd = Value.Date
                Me.CheckDirty()
            End Set
        End Property
        Public Property DateRecd() As Date
            Get
                Return dtDateRecd.Date
            End Get
            Set(ByVal Value As Date)
                dtDateRecd = Value.Date
                Me.CheckDirty()
            End Set
        End Property
        Public Property DateSigned() As Date
            Get
                Return dtDateSigned.Date
            End Get
            Set(ByVal Value As Date)
                dtDateSigned = Value.Date
                Me.CheckDirty()
            End Set
        End Property
        Public Property ALLDType() As Integer
            Get
                Return nALLDType
            End Get
            Set(ByVal Value As Integer)
                nALLDType = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property InertMaterial() As Integer
            Get
                Return nInertMaterial
            End Get
            Set(ByVal Value As Integer)
                nInertMaterial = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property LCPInstallDate() As Date
            Get
                Return dtLCPInstallDate.Date
            End Get
            Set(ByVal Value As Date)
                dtLCPInstallDate = Value.Date
                Me.CheckDirty()
            End Set
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
        Public Property LTTDate() As Date
            Get
                Return dtLTTDate.Date
            End Get
            Set(ByVal Value As Date)
                dtLTTDate = Value.Date
                Me.CheckDirty()
            End Set
        End Property
        Public Property PipeCPTest() As Date
            Get
                Return dtPipeCPTest.Date
            End Get
            Set(ByVal Value As Date)
                dtPipeCPTest = Value.Date
                Me.CheckDirty()
            End Set
        End Property
        Public Property DateShearTest() As Date
            Get
                Return dtShearTest.Date
            End Get
            Set(ByVal Value As Date)
                dtShearTest = Value.Date
                Me.CheckDirty()
            End Set
        End Property
        Public Property DatePipeSecInsp() As Date
            Get
                Return dtPipeSecInsp.Date
            End Get
            Set(ByVal Value As Date)
                dtPipeSecInsp = Value.Date
                Me.CheckDirty()
            End Set
        End Property
        Public Property DatePipeElecInsp() As Date
            Get
                Return dtPipeElecInsp.Date
            End Get
            Set(ByVal Value As Date)
                dtPipeElecInsp = Value.Date
                Me.CheckDirty()
            End Set
        End Property
        Public Property PipeCPType() As Integer
            Get
                Return nPipeCPType
            End Get
            Set(ByVal Value As Integer)
                nPipeCPType = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property PipeInstallDate() As Date
            Get
                Return dtPipeInstallDate.Date
            End Get
            Set(ByVal Value As Date)
                dtPipeInstallDate = Value.Date
                Me.CheckDirty()
            End Set
        End Property
        Public Property PipeLD() As Integer
            Get
                Return nPipeLD
            End Get
            Set(ByVal Value As Integer)
                nPipeLD = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property PipeManufacturer() As Integer
            Get
                Return nPipeManufacturer
            End Get
            Set(ByVal Value As Integer)
                nPipeManufacturer = Value
                Me.CheckDirty()
            End Set
        End Property
        Public ReadOnly Property PipeManufacturerOriginal() As Integer
            Get
                Return onPipeManufacturer
            End Get
        End Property
        Public ReadOnly Property PipeMatDescOriginal() As Integer
            Get
                Return onPipeMatDesc
            End Get
        End Property
        Public Property PipeMatDesc() As Integer
            Get
                Return nPipeMatDesc
            End Get
            Set(ByVal Value As Integer)
                nPipeMatDesc = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property PipeModDesc() As Integer
            Get
                Return nPipeModDesc
            End Get
            Set(ByVal Value As Integer)
                nPipeModDesc = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property PipeOtherMaterial() As String
            Get
                Return strPipeOtherMaterial
            End Get
            Set(ByVal Value As String)
                strPipeOtherMaterial = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property PipeStatusDesc() As Integer
            Get
                Return nPipeStatusDesc
            End Get
            Set(ByVal Value As Integer)
                nPipeStatusDesc = Value
                Me.CheckDirty()
            End Set
        End Property
        Public ReadOnly Property PipeStatusDescOriginal() As Integer
            Get
                Return onPipeStatusDesc
            End Get
        End Property
        Public Property PipeTypeDesc() As Integer
            Get
                Return nPipeTypeDesc
            End Get
            Set(ByVal Value As Integer)
                nPipeTypeDesc = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property PipingComments() As String
            Get
                Return strPipingComments
            End Get
            Set(ByVal Value As String)
                strPipingComments = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property PipeInstallationPlannedFor() As Date
            Get
                Return dtPipeInstallationPlannedFor.Date
            End Get
            Set(ByVal Value As Date)
                dtPipeInstallationPlannedFor = Value.Date
                Me.CheckDirty()
            End Set
        End Property
        Public Property PlacedInServiceDate() As Date
            Get
                Return dtPlacedInServiceDate.Date
            End Get
            Set(ByVal Value As Date)
                dtPlacedInServiceDate = Value.Date
                Me.CheckDirty()
            End Set
        End Property
        Public Property SubstanceComments() As Integer
            Get
                Return nSubstanceComments
            End Get
            Set(ByVal Value As Integer)
                nSubstanceComments = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property SubstanceDesc() As Integer
            Get
                Return nSubstanceDesc
            End Get
            Set(ByVal Value As Integer)
                nSubstanceDesc = Value
                'Me.CheckDirty()
            End Set
        End Property
        Public Property TermCPLastTested() As Date
            Get
                Return dtTermCPLastTested.Date
            End Get
            Set(ByVal Value As Date)
                dtTermCPLastTested = Value.Date
                Me.CheckDirty()
            End Set
        End Property
        Public Property TermCPTypeTank() As Integer
            Get
                Return nTermCPTypeTank
            End Get
            Set(ByVal Value As Integer)
                nTermCPTypeTank = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property TermCPTypeDisp() As Integer
            Get
                Return nTermCPTypeDisp
            End Get
            Set(ByVal Value As Integer)
                nTermCPTypeDisp = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property PipeCPInstalledDate() As Date
            Get
                Return dtPipeCPInstalledDate.Date
            End Get
            Set(ByVal Value As Date)
                dtPipeCPInstalledDate = Value.Date
                Me.CheckDirty()
            End Set
        End Property
        Public Property TermCPInstalledDate() As Date
            Get
                Return dtTermCPInstalledDate.Date
            End Get
            Set(ByVal Value As Date)
                dtTermCPInstalledDate = Value.Date
                Me.CheckDirty()
            End Set
        End Property
        Public Property TermTypeDisp() As Integer
            Get
                Return nTermTypeDisp
            End Get
            Set(ByVal Value As Integer)
                nTermTypeDisp = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property TermTypeTank() As Integer
            Get
                Return nTermTypeTank
            End Get
            Set(ByVal Value As Integer)
                nTermTypeTank = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property Deleted() As Boolean
            Get
                Return boolDeleted
            End Get
            Set(ByVal Value As Boolean)
                boolDeleted = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property IsDirty() As Boolean
            Get
                Return bolIsDirty
            End Get

            Set(ByVal value As Boolean)
                bolIsDirty = value
            End Set
        End Property
        Public Property CompartmentID() As String
            Get
                Return nTankID.ToString + "|" + nCompartmentNumber.ToString
            End Get
            Set(ByVal Value As String)
                Dim str() As String = Value.Split("|")
                TankID = CType(str(0), Integer)
                CompartmentNumber = CType(str(1), Integer)
                Me.CheckDirty()
            End Set
        End Property
        Public Property CompartmentNumber() As Integer
            Get
                Return nCompartmentNumber
            End Get
            Set(ByVal Value As Integer)
                nCompartmentNumber = Value
            End Set
        End Property
        Public Property CompartmentSubstance() As Integer
            Get
                Return nCompartmentSubstance
            End Get
            Set(ByVal Value As Integer)
                nCompartmentSubstance = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property CompartmentCERCLA() As Integer
            Get
                Return nCompartmentCERCLA
            End Get
            Set(ByVal Value As Integer)
                nCompartmentCERCLA = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property CompartmentFuelType() As Integer
            Get
                Return nCompartmentFuelType
            End Get
            Set(ByVal Value As Integer)
                nCompartmentFuelType = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property TankSiteID() As Integer
            Get
                Return nTankSiteID
            End Get
            Set(ByVal Value As Integer)
                nTankSiteID = Value
            End Set
        End Property
        Public Property AttachedPipeID() As Integer
            Get
                Return nAttachedPipeId
            End Get
            Set(ByVal Value As Integer)
                nAttachedPipeId = Value
            End Set
        End Property
        Public Property POU() As Boolean
            Get
                Return bolPOU
            End Get
            Set(ByVal Value As Boolean)
                bolPOU = Value
            End Set
        End Property
        Public Property NonPre88() As Boolean
            Get
                Return bolNonPre88
            End Get
            Set(ByVal Value As Boolean)
                bolNonPre88 = Value
            End Set
        End Property
        Public Property FacilityPowerOff() As Boolean
            Get
                Return bolFacPowerOff
            End Get
            Set(ByVal Value As Boolean)
                bolFacPowerOff = Value
            End Set
        End Property
        Public Property FacCapStatus() As Integer
            Get
                Return nFacCapStatus
            End Get
            Set(ByVal Value As Integer)
                nFacCapStatus = Value
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
